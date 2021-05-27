package com.grapecity.spreadjsworld.Controller;


import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;

@Controller
@RequestMapping({"/", "Home"})
public class IndexController {

    @RequestMapping("/")
    public String index(Model model) {
        return "index";
    }

}
